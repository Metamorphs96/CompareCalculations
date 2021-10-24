function StructuralCalcs100
    Cpe=-0.7
    qz=0.96			%kPa
    pn=Cpe*qz		%kPa
    s=3             %m
    w=pn*s			%kN/m
    L=24		    %m
    inc=L/12

    x=zeros(1,13)
    M=zeros(1,13)
    printf('Start Loop\n')
    for i=1:13
        if i==1
            x(i)=0
        else
            x(i)=x(i-1)+inc
        end    
        M(i)=((w*x(i))/2)*(L-x(i))
    end
    plot(x,M)
    title('Moment vs Span','fontsize',20);
    xlabel('Span [m]')
    ylabel('Moment [kNm]')
end

function BM_ss1=(w,x,L)
    BM_ss1=(w * x) / 2 * (L - x)
end